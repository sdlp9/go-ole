// +build windows

package ole

import (
	"unsafe"
)

func NewVariantFromSafeArrayFloat64(slice []float64) VARIANT {
	array, _ := safeArrayCreateVector(VT_R8, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []float64 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return NewVariant(VT_ARRAY|VT_R8, int64(uintptr(unsafe.Pointer(array))))
}

 